Attribute VB_Name = "Lot_fnBarras"
'---------------------------------------------------------------------------------------
' Module    : fn_Barras
' DateTime  : 27/05/2007 08:21
' Author    : Carlos Almela Baeza
' Purpose   : Aglutina el control del las macros
'---------------------------------------------------------------------------------------
Option Explicit
Public mobjEstadoAplicacion    As New EstadoAplicacion  'Objeto estado de la aplicaci�n
Private DB                     As New BdDatos           'Objeto Base de Datos
'---------------------------------------------------------------------------------------
' Procedimiento : Inicio_Aplicacion
' Creaci�n      : 01-nov-2006 17:57
' Autor         : Carlos Almela Baeza
' Objeto        : Configura la barra de herramientas del aplicativo
'---------------------------------------------------------------------------------------
'
Public Sub Inicio_Aplicacion()
    Application.ScreenUpdating = False                          'Desactiva el refresco de pantalla
    mobjEstadoAplicacion.OcultarTodasLasBarrasDeHerramientas    'Oculta todas las Barras de herramientas
    Application.Caption = "Hoja de Control de la Primi"         'Titulo del Libro
    Crear_Barra_Herramientas (BARRA_FUNCIONES)                  'Crea la Barra de herramientas particularizada de la aplicaci�n
    DB.Ir_A_Hoja "Movimientos"                                  'Posicionarse en el inicio
    Application.ScreenUpdating = True                           'Activa el refresco de pantalla
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Fin_Aplicacion
' Creaci�n      : 01-nov-2006 17:57
' Autor         : Carlos Almela Baeza
' Objeto        : Deja el Excel como estaba antes de iniciar la aplicaci�n
'---------------------------------------------------------------------------------------
'
Public Sub Fin_Aplicacion()
    Application.ScreenUpdating = False                           'Desactiva el refresco  de pantalla
    ActiveWindow.Caption = Empty                                 'Elimina el t�tulo de la ventana
    If mobjEstadoAplicacion.Existe_Barra(BARRA_FUNCIONES) Then   'Comprueba si est� la barra
        Application.CommandBars(BARRA_FUNCIONES).Delete          'Borra la barra de la aplicaci�n
    End If
    Application.ScreenUpdating = True                            'Activa el refresco de pantalla
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Borra_Barra_Herramientas
' Creaci�n      : 16-dic-2006 21:25
' Autor         : Carlos Almela Baeza
' Objeto        : Borrado de Barra de la aplicaci�n
'---------------------------------------------------------------------------------------
'
Public Sub Borra_Barra_Herramientas(Nombre_Barra As String)
    'Borra la barra de la aplicaci�n, si existe
    If mobjEstadoAplicacion.Existe_Barra(Nombre_Barra) Then
        Application.CommandBars(Nombre_Barra).Delete
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Crear_Barra_Herramientas
' Creaci�n      : 03-ago-2006 23:22
' Autor         : Carlos Almela Baeza
' Objeto        : Crea la barra de herramientas que invocar� a las macros de
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
            .Position = msoBarTop           'Posici�n Top de la barra
            .Visible = True                 'Visibilidad
        End With
        '.Style = msoButtonIconAndCaption        'Estilo Texto e Imagen
        'Orden de Funciones:
        ' 1.- Verificar:
        ' 2.- Colorear Resultados
        ' 3.- Simulaci�n
        ' Iconos:  2144  icono de las ruedas dentadas
        '          1664  Validar
        '           220  check box
        '           417  dibujo
        '          1691  bote pintura
        '           300  Estadisticas
        '            49  Interrogante
        '           341  Bombilla mas admiraci�n
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
                ' A�ade y crea un boton para la funci�n "Obtener estad�sticas"
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
                '  Obtenci�n de estadisticas de un n�mero
                '
                Set btnValidar = my_barra.Controls.Add(msoControlButton)
                With btnValidar
                    .Caption = "Estadisticas de un N�mero"
                    .Enabled = True
                    .Visible = True
                    .FaceId = 2147
                    .BeginGroup = True
                    .Style = msoButtonIconAndCaption
                    .OnAction = "btn_Prob_TiemposMedios"
                End With
                
                '
                '  Informaci�n de los resultados
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
                ' A�ade y crea un boton para la funci�n "M�todo �ptimo"
                '
'                Set btnValidar = my_barra.Controls.Add(msoControlButton)
'                With btnValidar
'                    .Caption = "M�todo Optimo"
'                    .Enabled = True
'                    .Visible = True
'                    .FaceId = 341
'                    .Style = msoButtonIconAndCaption
'                    .BeginGroup = True
'                    .OnAction = "btn_Metodo_Optimo"
'                End With
                '
                ' A�ade y crea un boton para la funci�n "Simulaci�n Varios M�todos"
                '
'                Set btnValidar = my_barra.Controls.Add(msoControlButton)
'                With btnValidar
'                    .Caption = "Simulaci�n Varios M�todos"
'                    .Enabled = True
'                    .Visible = True
'                    .FaceId = 156
'                    .Style = msoButtonIconAndCaption
'                    .BeginGroup = True
'                    .OnAction = "btn_SimularVariosMetodos"
'                End With
                '
                ' A�ade y crea un boton para la funci�n "Calcular la Sugerencia"
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
                ' A�ade y crea un boton para la funci�n "Comprobar Metodo"
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
                ' A�ade y crea un boton para la funci�n "Verificar Pronosticos"
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
                ' A�ade y crea un boton para la funci�n "Versi�n de aplicaci�n"
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
    Dim DB          As New BdDatos
    Dim objGraf     As Shape
    DB.Ir_A_Hoja ("Salida")                    ' selecciona la hoja que contendr� la salida
    If (ActiveSheet.Shapes.Count > 0) Then     ' Si existe alg�n gr�fico
       For Each objGraf In ActiveSheet.Shapes  ' Para cada Gr�fico en la hoja activa
            objGraf.Delete                     ' Elimina el gr�fico
       Next objGraf
    End If
    Cells.Select                               ' Selecciona todo el contenido
    Selection.ColumnWidth = 10                 ' Establece el ancho de las columnas a 10 puntos
    Selection.Clear                            ' Borra la selcci�n, el contenido y los formatos
    Selection.ClearComments                    ' Borra los comentarios
End Sub
