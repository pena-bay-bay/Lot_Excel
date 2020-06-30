Attribute VB_Name = "Lot_Funciones"
' *============================================================================*
' *
' *     Fichero    : Lot_Funciones.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 24/jun/2020 18:11:09
' *     Versión    : 1.1
' *     Propósito  : Recopilar funciones comunes
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public mobjEstadoAplicacion    As New EstadoAplicacion  'Objeto estado de la aplicación

'--- Métodos Privados ---------------------------------------------------------*
'--- Métodos Públicos ---------------------------------------------------------*

'---------------------------------------------------------------------------------------
' Procedimiento : Colorea_Celda
' Creación      : 16-dic-2006 21:14
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea la celda según el número y la muestra a aplicar
'---------------------------------------------------------------------------------------
'
Public Sub Colorea_Celda(celda As Range, _
                          Numero As Variant, _
                          ByRef objMuestra As Muestra, _
                          objMetodo As Metodo)
    Dim i        As Integer                     'Número en formato Entero
    Dim m_iColor As Integer                     'Color de la celda
   On Error GoTo Colorea_Celda_Error
    i = CInt(Numero)                            'Obtiene el entero del número
    m_iColor = xlNone                           'Inicializa el color a automático
    
    'Selección de la matriz de números según el método
    
    
    Select Case (objMetodo.CriteriosOrdenacion)

'        Case ordSinDefinir

'        Case ordProbabilidad

        Case ordProbTiempoMedio:
            m_iColor = get_color_array(objMuestra.Matriz_ProbTiempos, i)

        Case ordFrecuencia:
            m_iColor = get_color_array(objMuestra.Matriz_ProbFrecuencias, i)
        
'        Case ordAusencia

'        Case ordTiempoMedio

'        Case ordDesviacion

'        Case ordProximaFecha

'        Case ordModa

        Case Else: m_iColor = get_color_array(objMuestra.Matriz_Probabilidades, i)
    End Select
    
    celda.Value = Numero                       'Asigna el Numero a la celda
    DestacarRango celda, m_iColor              'Colorea la celda

   On Error GoTo 0
   Exit Sub

Colorea_Celda_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Colorea_Celda of Módulo Lot_VerSorteos"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : Colorea_Celda
' Creación      : 16-dic-2006 21:14
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea la celda según el número y la muestra a aplicar
'---------------------------------------------------------------------------------------
'
Public Sub Colorea_CeldaProb(celda As Range, _
                          Numero As Variant, _
                          ByRef objMuestra As Muestra)
                          
    Dim i        As Integer                     'Número en formato Entero
    Dim m_iColor As Integer                     'Color de la celda
   On Error GoTo Colorea_Celda_Error
    i = CInt(Numero)                            'Obtiene el entero del número
    m_iColor = xlNone                           'Inicializa el color a automático
    
    m_iColor = get_color_array(objMuestra.Matriz_Probabilidades, i)
    
    celda.Value = Numero                       'Asigna el Numero a la celda
    DestacarRango celda, m_iColor              'Colorea la celda

   On Error GoTo 0
   Exit Sub

Colorea_Celda_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Colorea_Celda of Módulo Lot_VerSorteos"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : ColoreaCelda
' Creación      : 09-nov-2006 10:09
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea una celda o rango de celdas
'---------------------------------------------------------------------------------------
'
Public Sub ColoreaCelda(cldTarget As Range, color As Integer)
    If (color = COLOR_TERMINACION0) Then
        cldTarget.Font.ColorIndex = COLOR_NUMCOMPLE
    Else
        cldTarget.Font.ColorIndex = xlColorIndexAutomatic
    End If
    cldTarget.Interior.ColorIndex = color
End Sub

'
'
'
Public Function GetFechaRegistro(datRegistro As Integer, Optional datJuego As Juego = PrimitivaBonoloto)
    Dim maxRegistro     As Integer
    Dim maxFecha        As Date
    Dim m_dtFecha       As Date
    Dim mDB             As BdDatos
    '
    ' Obtengo los datos del último registro
    '
    Set mDB = New BdDatos
    
    maxFecha = mDB.UltimoResultado
    maxRegistro = mDB.UltimoRegistro
    '
    ' Si es inferior al máximo
    '
    Select Case datRegistro
    Case Is = maxRegistro
        m_dtFecha = maxFecha
    Case Else
        m_dtFecha = mDB.GetSimulacionFecha(datRegistro)
    End Select
    GetFechaRegistro = m_dtFecha
End Function


'
'
'
Public Function GetRegistroFecha(datFecha As Date, Optional datJuego As Juego = PrimitivaBonoloto)
    Dim mRes            As Sorteo
    Dim maxRegistro     As Integer
    Dim maxFecha        As Date
    Dim mIDifDias       As Integer
    Dim iDiaSem         As Integer
    Dim mDB             As BdDatos
    
    
    Set mDB = New BdDatos
    
    maxFecha = mDB.UltimoResultado
    maxRegistro = mDB.UltimoRegistro
    
    If (datFecha < maxFecha) Then
        mRes = mDB.Get_Resultado(datFecha)
        GetRegistroFecha = mRes
    Else
        mIDifDias = datFecha - maxFecha
    End If
    iDiaSem = Weekday(datFecha, vbMonday)
    Select Case datJuego
    ' L, M, X, J, V, S
    Case PrimitivaBonoloto
    ' L, M, X, V
    Case Juego.Bonoloto
    ' J, S
    Case Juego.LoteriaPrimitiva
    ' M, V
    Case Juego.Euromillones
    ' D
    Case Juego.GordoPrimitiva
    
    End Select

End Function

'
'
'
Public Function GetModa(datValores As Variant) As Double
    Dim mResult As Double
    On Error Resume Next
    mResult = Application.WorksheetFunction.Mode(datValores)
    If Err.Number <> 0 Then
        mResult = Application.WorksheetFunction.Median(datValores)
    End If
    GetModa = mResult
End Function

'---------------------------------------------------------------------------------------
' Procedure : Version_Libreria
' Author    : CHARLY
' Date      : sáb, 14/01/2012 19:50
' Purpose   : Visualiza la version de las macros
'---------------------------------------------------------------------------------------
'
Public Sub Version_Libreria()
    Dim Version As String
    Version = " La versión de la librería es la:" & vbCrLf _
              & vbTab & lotVersion & vbCrLf _
              & "de fecha " & vbTab & lotFeVersion
    MsgBox Version, vbApplicationModal + vbInformation + vbOKOnly, "Librería de Funciones de la Loteria"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CALCULOON
' Author    : CAB3780Y
' Date      : 04/08/2009
' Purpose   : Activa el refresco de hojas
'---------------------------------------------------------------------------------------
Public Sub CALCULOON()
    With Application
        .Calculation = xlAutomatic                          'cálculo automático
        .MaxChange = 0.001
        .CalculateBeforeSave = False
        .ScreenUpdating = True                              'actualizar pantalla
        .ErrorCheckingOptions.BackgroundChecking = False    'no verificar errores formulas
    End With
    With ActiveWorkbook
        .UpdateRemoteReferences = False                     'no actualizar ref. remotas
        .PrecisionAsDisplayed = False
        .SaveLinkValues = False
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CALCULOOFF
' Author    : CAB3780Y
' Date      : 04/08/2009
' Purpose   : Desactiva el refresco de hojas
'---------------------------------------------------------------------------------------
'
Public Sub CALCULOOFF()
    With Application
        .ScreenUpdating = False
        .Calculation = xlManual
        .MaxChange = 0.001
        .CalculateBeforeSave = False
    End With
    
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
        .SaveLinkValues = False
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : ExistenElem
' Creación      : 12-nov-2006 00:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function ExistenElem(col As Collection, Clave As String) As Boolean
    Dim prueba As Variant
    On Error Resume Next
    prueba = col.Item(Clave)
    ExistenElem = (Err <> 5)
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : EliminarTodosElementos
' Creación      : 12-nov-2006 00:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub EliminarTodosElementos(col As Collection)
    Do While col.Count
        col.Remove 1
    Loop
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : SustituirElem
' Creación      : 12-nov-2006 00:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub SustituirElem(col As Collection, indice As Variant, nuevoValor As Variant)
    col.Remove indice
    If VarType(indice) = vbString Then
        col.Add nuevoValor, indice
    Else
        If indice > col.Count Then
            col.Add nuevoValor, , col.Count
        Else
            col.Add nuevoValor, , indice
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : ShellSortAny
' Creación      : 12-nov-2006 00:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
'Public Sub ShellSortAny(arr As Variant, numEls As Long, descendente As Boolean)
'    Dim indice As Long, indice2 As Long, primerElem As Long
'    Dim distancia As Long, Valor As Variant
'
'   On Error GoTo ShellSortAny_Error
'
'    ' salir si no es un array
'    If VarType(arr) < vbArray Then Exit Sub
'
'    primerElem = LBound(arr)
'
'    ' encontrar el mejor valor para distancia
'    Do
'        distancia = distancia * 3 + 1
'    Loop Until distancia > numEls
'
'    ' ordenar el array
'    Do
'        distancia = distancia / 3
'        For indice = distancia + primerElem To numEls + primerElem - 1
'            Valor = arr(indice)
'            indice2 = indice
'            Do While (arr(indice2 - distancia) > Valor) Xor descendente
'                arr(indice2) = arr(indice2 - distancia)
'                indice2 = indice2 - distancia
'                If indice2 - distancia < primerElem Then Exit Do
'            Loop
'            arr(indice2) = Valor
'        Next
'    Loop Until distancia = 3
'
'   On Error GoTo 0
'   Exit Sub
'
'ShellSortAny_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & _
'        ") in procedure ShellSortAny of Módulo fn_collections"
'End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Ordenar
' Creación      : 30-Oct-2002 00:40
' Autor         : Carlos Almela Baeza
' Objeto        : Ordenar una matriz de una sola dimensión
'---------------------------------------------------------------------------------------
'
Public Sub Ordenar(ByRef matriz As Variant, _
                    Optional Ascendente As Boolean = True)
    
    Dim tsOrdenado  As Boolean          'indicador de matriz ordenada
    Dim MxLimite    As Integer          'Limite máximo del bucle
    Dim TmpDato     As Variant          'Dato temporal para el intercambio
    Dim i           As Integer          'Indice
    
  On Error GoTo Ordenar_Error:
  
    'Si no es un elemento array sale de la rutina
    If Not IsArray(matriz) Then Exit Sub
    
    'obtiene el número máximo de elementos de la matriz
    MxLimite = UBound(matriz)
    
    'Bucle de ordenación, se realiza hasta que esté ordenada
    Do
        tsOrdenado = True               'Se parte de matriz ordenada
        For i = 0 To MxLimite - 1       'Se revisa cada elemento con el siguiente
            If Ascendente Then
                If matriz(i) > matriz(i + 1) _
                And (matriz(i + 1) <> 0) Then
                    tsOrdenado = False
                    TmpDato = matriz(i)         'Guardamos la posicion iesima
                    matriz(i) = matriz(i + 1)   'pasamos la posision siguiente a la
                                                'iesima
                    matriz(i + 1) = TmpDato     'pasamos el dato guardado a la siguiente
                End If
            Else
                If matriz(i) < matriz(i + 1) _
                And (matriz(i + 1) <> 0) Then
                    tsOrdenado = False
                    TmpDato = matriz(i)         'Guardamos la posicion iesima
                    matriz(i) = matriz(i + 1)   'pasamos la posision siguiente a la
                                                'iesima
                    matriz(i + 1) = TmpDato     'pasamos el dato guardado a la siguiente
                End If
            End If
        Next i
    Loop Until tsOrdenado

  On Error GoTo 0
    Exit Sub
Ordenar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_fncollections.Ordenar", ErrSource)
    Err.Raise ErrNumber, "Lot_fncollections.Ordenar", ErrDescription
End Sub



'------------------------------------------------------------------------------*
'Función     : Ordenar2
'Fecha       : 28-Nov-1999
'Parametros  : Matriz de dos dimensiones
'Descripción : Ordena la matriz de paso
'                     x ->
'              matriz (0, 0) (0, 1)
'             y  |    (1, 0) (1, 1)
'                v    (2, 0) (2, 1)
'------------------------------------------------------------------------------*
Public Sub Ordenar2(ByRef matriz As Variant, _
                    Optional columna As Integer = 2, _
                    Optional Ascendente As Boolean = True)
                    
    Dim tsOrdenado          As Boolean
    Dim limiteY             As Integer
    Dim limiteX             As Integer
    Dim TmpDato()           As Variant
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    
   On Error GoTo Ordenar2_Error

    limiteY = UBound(matriz, 1)
    limiteX = UBound(matriz, 2)
    ReDim TmpDato(limiteX)
    
    j = IIf(columna = 1, 0, 1)
    
    If Not IsArray(matriz) Then Exit Sub
    
    Do
        tsOrdenado = True
        For i = 0 To limiteY - 1
            If Ascendente Then
                If (matriz(i, j) > matriz(i + 1, j)) Then
                    tsOrdenado = False
                    For k = 0 To limiteX
                        TmpDato(k) = matriz(i, k)
                        matriz(i, k) = matriz(i + 1, k)
                        matriz(i + 1, k) = TmpDato(k)
                    Next k
                End If
            Else
                If (matriz(i, j) < matriz(i + 1, j)) Then
                    tsOrdenado = False
                    For k = 0 To limiteX
                        TmpDato(k) = matriz(i, k)
                        matriz(i, k) = matriz(i + 1, k)
                        matriz(i + 1, k) = TmpDato(k)
                    Next k
                End If
            End If
        Next i
    Loop Until tsOrdenado

   On Error GoTo 0
   Exit Sub

Ordenar2_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_fncollections.Ordenar2", ErrSource)
    Err.Raise ErrNumber, "Lot_fncollections.Ordenar2", ErrDescription
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : Inicio_Aplicacion
' Creación      : 01-nov-2006 17:57
' Autor         : Carlos Almela Baeza
' Objeto        : Configura la barra de herramientas del aplicativo
'---------------------------------------------------------------------------------------
'
Public Sub Inicio_Aplicacion()
    Application.ScreenUpdating = False                          'Desactiva el refresco de pantalla
    mobjEstadoAplicacion.OcultarTodasLasBarrasDeHerramientas    'Oculta todas las Barras de herramientas
    Application.Caption = "Hoja de Control de la Primi"         'Titulo del Libro
    Crear_Barra_Herramientas (BARRA_FUNCIONES)                  'Crea la Barra de herramientas particularizada de la aplicación
    Ir_A_Hoja "Resultados"                                       'Posicionarse en el inicio
    Application.ScreenUpdating = True                           'Activa el refresco de pantalla
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Fin_Aplicacion
' Creación      : 01-nov-2006 17:57
' Autor         : Carlos Almela Baeza
' Objeto        : Deja el Excel como estaba antes de iniciar la aplicación
'---------------------------------------------------------------------------------------
'
Public Sub Fin_Aplicacion()
    Application.ScreenUpdating = False                           'Desactiva el refresco  de pantalla
    ActiveWindow.Caption = Empty                                 'Elimina el título de la ventana
    If mobjEstadoAplicacion.Existe_Barra(BARRA_FUNCIONES) Then   'Comprueba si está la barra
        Application.CommandBars(BARRA_FUNCIONES).Delete          'Borra la barra de la aplicación
    End If
    Application.ScreenUpdating = True                            'Activa el refresco de pantalla
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Borra_Barra_Herramientas
' Creación      : 16-dic-2006 21:25
' Autor         : Carlos Almela Baeza
' Objeto        : Borrado de Barra de la aplicación
'---------------------------------------------------------------------------------------
'
Public Sub Borra_Barra_Herramientas(Nombre_Barra As String)
    'Borra la barra de la aplicación, si existe
    If mobjEstadoAplicacion.Existe_Barra(Nombre_Barra) Then
        Application.CommandBars(Nombre_Barra).Delete
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Crear_Barra_Herramientas
' Creación      : 03-ago-2006 23:22
' Autor         : Carlos Almela Baeza
' Objeto        : Crea la barra de herramientas que invocará a las macros de
'                 la hoja, o a la funcionalidad
'---------------------------------------------------------------------------------------
'
Public Sub Crear_Barra_Herramientas(Nombre_Barra As String)
    Dim btnValidar  As CommandBarButton     'Objeto Boton de la barra de herramientas
    Dim my_barra    As CommandBar           'Barra de Herramientas
   
    'Consulta si la barra del programa ya existe, y si no existe crea una
   On Error GoTo Crear_Barra_Herramientas_Error
            
    If Not mobjEstadoAplicacion.Existe_Barra(Nombre_Barra) Then
        
        'Crea una barra de herramientas especial para esta hoja
        Set my_barra = Application.CommandBars.Add(Nombre_Barra, , False, True)
        With my_barra
            .Position = msoBarTop           'Posición Top de la barra
            .Visible = True                 'Visibilidad
        End With
        '.Style = msoButtonIconAndCaption        'Estilo Texto e Imagen
        'Orden de Funciones:
        ' 1.- Verificar:
        ' 2.- Colorear Resultados
        ' 3.- Simulación
        ' Iconos:  2144  icono de las ruedas dentadas
        '          1664  Validar
        '           220  check box
        '           417  dibujo
        '          1691  bote pintura
        '           300  Estadisticas
        '            49  Interrogante
        '           341  Bombilla mas admiración
        '           156  play
        Select Case (Nombre_Barra)
            Case BARRA_FUNCIONES
                
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Combrobar Boletos"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 1664              'Validar
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_ComprobarBoletos"
                End With
                '
                ' --------------    Boton de Colorear resultados
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Colorear Sorteos"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 1691              'Bote Pintura
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_Colorear"
                End With
                
                '
                ' Grupo Colorear resultados
                ' Añade y crea un boton para la función "Obtener estadísticas"
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Obtener Estadisticas"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 2140                          'Porcentajes
                    .Style = msoButtonIconAndCaption
                    .BeginGroup = True
                    .OnAction = "btn_Obtener_Estadisticas"
                End With
                
                '
                '  Obtención de estadisticas de un número
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Estadisticas de un Número"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 2147
                    .BeginGroup = True
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_Prob_TiemposMedios"
                End With
                
                '
                '  Información de los resultados
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Caracteristicas de Resultados"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 2144
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_VerificarSorteos"
                End With
                
                '
                ' Añade y crea un boton para la función "Método óptimo"
                '
'                Set btnValidar = my_barra.Controls.Add(msoControlButton)
'                With btnValidar
'                    .Caption = "Método Optimo"
'                    .Enabled = True
'                    .Visible = True
'                    .FaceId = 341
'                    .Style = msoButtonIconAndCaption
'                    .BeginGroup = True
'                    .OnAction = "btn_Metodo_Optimo"
'                End With
                '
                ' Añade y crea un boton para la función "Simulación Varios Métodos"
                '
'                Set btnValidar = my_barra.Controls.Add(msoControlButton)
'                With btnValidar
'                    .Caption = "Simulación Varios Métodos"
'                    .Enabled = True
'                    .Visible = True
'                    .FaceId = 156
'                    .Style = msoButtonIconAndCaption
'                    .BeginGroup = True
'                    .OnAction = "btn_SimularVariosMetodos"
'                End With
                '
                ' Añade y crea un boton para la función "Calcular la Sugerencia"
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Sugerencias"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 341
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_SugerirApuestas"
                End With
                '
                ' Añade y crea un boton para la función "Comprobar Metodo"
                '
'                Set btnValidar = my_barra.Controls.Add(msoControlButton)
'                With btnValidar
'                    .Caption = "Comprobar Metodo"
'                    .Enabled = True
'                    .Visible = True
'                    .FaceId = 2144
'                    .Style = msoButtonIconAndCaption
'                    .OnAction = "btn_ComprobarMetodo"
'                End With
                '
                ' Añade y crea un boton para la función "Verificar Pronosticos"
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Comprobar Apuestas"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 1664              'Validar
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_ComprobarApuestas"
                End With
                '
                ' Añade y crea un boton para la función "Versión de aplicación"
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Version"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 49
                    .Style = msoButtonIconAndCaption
                    .OnAction = "Version_Libreria"
                End With
        
        End Select
    Else
        'Si ya existe la barra la visualiza
        Application.CommandBars(Nombre_Barra).Visible = True
    End If
            
   On Error GoTo 0
       Exit Sub
            
Crear_Barra_Herramientas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_fnBarras.Crear_Barra_Herramientas", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "Lot_fnBarras.Crear_Barra_Herramientas", ErrDescription

End Sub

Public Sub Borra_Salida()
    Dim objGraf     As Shape
    
    Ir_A_Hoja ("Salida")                    ' selecciona la hoja que contendrá la salida
    If (ActiveSheet.Shapes.Count > 0) Then     ' Si existe algún gráfico
       For Each objGraf In ActiveSheet.Shapes  ' Para cada Gráfico en la hoja activa
            objGraf.Delete                     ' Elimina el gráfico
       Next objGraf
    End If
    Cells.Select                               ' Selecciona todo el contenido
    Selection.ColumnWidth = 10                 ' Establece el ancho de las columnas a 10 puntos
    Selection.Clear                            ' Borra la selcción, el contenido y los formatos
    Selection.ClearComments                    ' Borra los comentarios
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Ir_A_Hoja
' Creación      : 05-nov-2006 20:38
' Autor         : Carlos Almela Baeza
' Objeto        : Selecciona la hoja del libro donde actua la macro
'---------------------------------------------------------------------------------------
'
Public Sub Ir_A_Hoja(hoja As String)
    Dim Wrk As Workbook
    Dim Pagina As Worksheet
    For Each Wrk In Workbooks
        For Each Pagina In Wrk.Sheets
            If Pagina.Name = hoja Then
                Pagina.Select
                Exit Sub
            End If
        Next Pagina
    Next Wrk
End Sub
' *===========(EOF):
