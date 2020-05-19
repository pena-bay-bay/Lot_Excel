VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetodoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : MetodoController.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : lu., 30/mar/2020 20:57:41
' *     Versión    : 1.0
' *     Propósito  : Controlador de Metodo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mModel          As MetodoModel                  ' Modelo de negocio
Private frmEditView     As frmMetodoEditView            ' Formulario de edición


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mModel = New MetodoModel
End Sub



Private Sub Class_Terminate()
    Set mModel = Nothing
End Sub


'--- Métodos Públicos ---------------------------------------------------------*

Public Sub SetLinePerPage(datLineas As Integer)
    mModel.LinePerPage = datLineas
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : InitSelectView
' Fecha          : ma., 24/mar/2020 19:54:32
' Propósito      : Inicializa la vista con datos de la tabla
' Parámetros     : Formulario de visualización
'------------------------------------------------------------------------------*
Public Sub InitSelectView(datForm As Variant)
 On Error GoTo InitSelectView_Error
    '
    '   Inicializamos el numero de páginas
    '
    mModel.LinePerPage = 6
    '
    '   Buscamos la primera página
    '
    mModel.Search (1)
    '
    '   Inicializamos los controles del formulario
    '
    datForm.InitForm
    '
    '   Actualizamos datos del formulario
    '
    datForm.Refresh mModel
    '
    '
    '
  On Error GoTo 0
    Exit Sub
InitSelectView_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoController.InitSelectView", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : InitEditView
' Fecha          : lu., 30/mar/2020 20:57:41
' Propósito      : Inicializa la vista con datos de la tabla
' Parámetros     : Identificador de edición
'------------------------------------------------------------------------------*
Public Function InitEditView(datId As Integer, datPronosticos As Integer) As MetodoModel
    Dim mModoAlta As Boolean
  On Error GoTo InitEditView_Error
    '
    '   Deducimos el modo de funcionamiento del formulario
    '
    If datId = 0 Then
        mModoAlta = True
        mModel.SetNew
        mModel.Metodo.Pronosticos = datPronosticos
    Else
        mModoAlta = False
        '
        '   Si tenemos Id
        '
        mModel.GetRecord (datId)
    End If
    '
    '   Crea un formulario de edición de métodos
    '
    Set frmEditView = New frmMetodoEditView
    '
    '   inicializa el estado
    '
    frmEditView.Tag = ESTADO_INICIAL
    '
    '   Actualiza formulario con el modelo
    '
    frmEditView.ModoAlta = mModoAlta
    frmEditView.Refresh mModel
    '
    '   Selecciona parametros del proceso
    '
    Do While frmEditView.Tag <> BOTON_CERRAR
    
        ' Se inicializa el boton cerrar para salir del bucle
        frmEditView.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        frmEditView.Show
       
        'Se bifurca la función
        Select Case frmEditView.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                frmEditView.Tag = BOTON_CERRAR
            
            Case EJECUTAR          ' Se ha pulsado el botón ejecutar
                If mModoAlta Then
                    '
                    '   Agregamos el current  método
                    '
                    mModel.Add
                Else
                    '
                    '  Guardamos el current método
                    '
                    mModel.Save
                End If
                frmEditView.Tag = BOTON_CERRAR
                
            Case BORRAR           ' Se ha pulsado el botón borrar
                '
                '   Borramos el current Metodo
                '
                mModel.Del mModel.Metodo.Id
                
                frmEditView.Tag = BOTON_CERRAR
        End Select
    Loop
    If mModoAlta Then
        '
        '   Si agregamos un registro vamos al final
        '
        mModel.Search mModel.TotalPages
    Else
        '
        '   Actualizamos la página actual
        '
        mModel.Search mModel.CurrentPage
    End If
    '
    '   Devolvemos el modelo para actualizar la vista
    '
    Set InitEditView = mModel
    
  On Error GoTo 0
InitEditView__CleanExit:
    Exit Function
InitEditView_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoController.InitEditView", ErrSource)
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : GoPageNumber
' Fecha          : ma., 21/abr/2020 17:59:06
' Propósito      : Visualizar la página n esima de la base de datos
' Parámetros     : Número de página a visualizar
'------------------------------------------------------------------------------*
Public Function GoPageNumber(datPage As Integer) As MetodoModel
  On Error GoTo GoPageNumber_Error
    '
    '   Busca la página solicitada
    '
    mModel.Search datPage
    '
    '  Devolvemos el modelo
    '
    Set GoPageNumber = mModel
  On Error GoTo 0
GoPageNumber__CleanExit:
    Exit Function
GoPageNumber_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoController.GoPageNumber", ErrSource)
End Function




'------------------------------------------------------------------------------*
' Procedimiento  : RefreshSelect
' Fecha          : ma., 31/mar/2020 19:14:10
' Propósito      : Actualiza el modelo en el formulario
' Parámetros     : Formulario a refrescar datos
'------------------------------------------------------------------------------*
Public Sub RefreshSelect(datForm As Variant)
    '
    '   Actualizamos datos del formulario
    '
    datForm.Refresh mModel
End Sub

' *===========(EOF): MetodoController.cls


