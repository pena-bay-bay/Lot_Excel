VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalTiempoMedio 
   Caption         =   "Estadisticas de un Número"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   OleObjectBlob   =   "frmCalTiempoMedio.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCalTiempoMedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'--------------------------------------------------------------------------------------*
' Module    : frmCalTiempoMedio.frm
' DateTime  : 10/02/2018 00:23
' Author    : Carlos Almela Baeza
' Purpose   : Obtener las estadisticas detalladas de uno o varios Numeros
'--------------------------------------------------------------------------------------*
Option Explicit
'
'-----------------------------|Variables|----------------------------------------------*
'
Private mPeriodo As New Periodo        ' Objeto que facilita el manejo de dos fechas
Private mMuestra As New Muestra        ' Objeto muestra
Private mParMuestra As New ParametrosMuestra ' Parametros de la muestra
Private mFechaSorteo As Date           ' Fecha de Sorteo o pronóstico
Private mNumero1 As Integer            ' Numero1 a evaluar
Private mNumero2 As Integer            ' Numero2 a evaluar
Private mNumero3 As Integer            ' Numero3 a evaluar
Private mCombinacion As Combinacion    ' Combinación de números
Private mDB As New BdDatos             ' Base de datos
Private mIniDataBase As Date           ' Fecha inicial de la base de datos
Private mFinDataBase As Date           ' Fecha final de la base de datos
Private mTipoProceso As Integer        ' Tipo de selección: 1- Todos, 2-Sorteo, 3-Numeros
Private mErrorValidacion As Long       ' Control de errores de validación
'
'   Validaciones de la clase
'
Private Const ERR_NumeroNOVALIDO            As Long = 1
Private Const ERR_NumeroNOEXISTE            As Long = 2
Private Const ERR_FECHANOVALIDA             As Long = 4
Private Const ERR_FECHANOSORTEO             As Long = 8
Private Const ERR_PARMUESTRAMAL             As Long = 16
Private Const ERR_NOESNumero                As Long = 32
Private Const ERR_FEININOVALIDA             As Long = 64
Private Const ERR_FEFINNOVALIDA             As Long = 128
'
'   Mensajes de validación
'
Private Const MSG_ERRORESVALIDACION         As String = "Los datos del formulario no cumplen las siguientes validaciones:" & vbCrLf
Private Const MSG_NumeroNOVALIDO            As String = "* El número debe estar comprendido entre 1 y 49."
Private Const MSG_NumeroNOEXISTE            As String = "* Debe introducir al menos un número."
Private Const MSG_FECHANOSORTEO             As String = "* La fecha de Sorteo/Análisis no es fecha de sorteo."
Private Const MSG_FECHANOVALIDA             As String = "* La fecha de Sorteo/Análisis no es valida."
Private Const MSG_NOESUNNumero              As String = "* Al menos hay un texto no numérico."
Private Const MSG_PARMUESTRAMAL             As String = "* Los parámetros de la muestra son erroneos..."
Private Const MSG_FEININOVALIDA             As String = "* La fecha inicial del período no es válida."
Private Const MSG_FEFINNOVALIDA             As String = "* La fecha final del período no es válida."
'
'
'-----------------------------|Propiedades|--------------------------------------------*
'
'
'--------------------------------------------------------------------------------------*
' Property  : MuestraCalculo
' DateTime  : 10/02/2018 00:23
' Author    : Carlos Almela Baeza
' Purpose   : Muestra de análisis estadistico
'--------------------------------------------------------------------------------------*
'
Public Property Get MuestraCalculo() As Muestra
    Set MuestraCalculo = mMuestra
End Property
'--------------------------------------------------------------------------------------*
' Property  : FechaSorteo
' DateTime  : 10/02/2018 00:23
' Author    : Carlos Almela Baeza
' Purpose   : Fecha de sorteo o previsión a analizar
'--------------------------------------------------------------------------------------*
'
Public Property Get FechaSorteo() As Date
    FechaSorteo = mFechaSorteo
End Property
'--------------------------------------------------------------------------------------*
' Property  : Combinacion
' DateTime  : 10/02/2018 00:23
' Author    : Carlos Almela Baeza
' Purpose   : Combinación de Numeros
'--------------------------------------------------------------------------------------*
'
Public Property Get Combinacion() As Combinacion
    Set Combinacion = mCombinacion
End Property
'--------------------------------------------------------------------------------------*
' Property  : TipoProceso
' DateTime  : 10/02/2018 00:23
' Author    : Carlos Almela Baeza
' Purpose   : Tipo de proceso a ejecutar
'--------------------------------------------------------------------------------------*
'
Public Property Get TipoProceso() As Integer
    TipoProceso = mTipoProceso
End Property
'
'
'-----------------------------|Eventos|------------------------------------------------*
'
'
'--------------------------------------------------------------------------------------*
' Procedure : UserForm_Initialize
' DateTime  : 10/02/2018 00:23
' Author    : Carlos Almela Baeza
' Purpose   : Inicializa el formulario
'--------------------------------------------------------------------------------------*
'
Private Sub UserForm_Initialize()

   On Error GoTo UserForm_Initialize_Error
    '
    '   Inicialización de objetos y variables
    '
    mFechaSorteo = Date
    Set mMuestra = New Muestra
    Set mCombinacion = New Combinacion
    Set mPeriodo = New Periodo
    '
    '   Obtenemos el periodo de fechas de la Base de datos
    '
    mIniDataBase = mDB.PrimerResultado
    mFinDataBase = mDB.UltimoResultado
    '
    '   Asignamos valores por defecto a las variables internas
    '
    mTipoProceso = 1    ' Todos los números por defecto
    '
    '   Cargamos el Combo con los criterios de periodos
    '
    mPeriodo.CargaTabla cboPeriodo
    '
    '   Seleccionamos un trimestre de datos
    '
    cboPeriodo.ListIndex = ctUltimoTrimestre
    mPeriodo.Tipo_Fecha = ctUltimoTrimestre
    '
    '   Visualizamos la información en los controles
    '
    VisualizaControles
   On Error GoTo 0
   Exit Sub

UserForm_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "frmCalTiempoMedio.UserForm_Initialize")
       '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCancelar_Click
' DateTime  : 02/05/2007 23:26
' Author    : Carlos Almela Baeza
' Purpose   : Cancela la presentación del formulario
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR               ' Asigna en la etiqueta del formulario
                                        ' la clave de cerrar
    Me.Hide                             ' Oculta el formulario
End Sub
Private Sub chkNumeros_Click()
    If chkNumeros.Value Then
        mTipoProceso = 3
        VisualizaControles
    End If
End Sub

Private Sub chkSorteo_Click()
    If chkSorteo.Value Then
        mTipoProceso = 2
        VisualizaControles
    End If
End Sub

Private Sub chkTodosNumeros_Click()
    If chkTodosNumeros.Value Then
        mTipoProceso = 1
        VisualizaControles
    End If
End Sub

Private Sub txtFePronostico_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(txtFePronostico.Text) Then
        mFechaSorteo = CDate(txtFePronostico.Text)
    End If
    '
    '   Visualizamos controles
    '
    VisualizaControles
End Sub


Private Sub cboPeriodo_Change()
    mPeriodo.Tipo_Fecha = cboPeriodo.ListIndex
    If mPeriodo.Tipo_Fecha <> 0 Then                ' Si el periodo no es personalizado
                                                    ' Si la fecha inicial es menor
                                                    ' que la base de datos
        If (mPeriodo.FechaInicial < mIniDataBase) Then
                                                    ' Selecciona la fecha menor de la
                                                    ' base de datos
            mPeriodo.FechaInicial = mIniDataBase
        End If
                                                    ' Si la fecha final es superior a
                                                    ' la fecha de la base de datos
        If (mPeriodo.FechaFinal > mFinDataBase) Then
                                                    ' Selecciona la fecha mayor de la base
                                                    ' de datos
            mPeriodo.FechaFinal = mFinDataBase
        End If
    End If
    '
    '
    '
    VisualizaControles
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cmdEjecutar_Click
' DateTime  : 02/05/2007 23:27
' Author    : Carlos Almela Baeza
' Purpose   : Validar parametros y salir del formulario
'---------------------------------------------------------------------------------------
'
Private Sub cmdEjecutar_Click()
    Dim strMensaje As String
    
 On Error GoTo cmdEjecutar_Click_Error
    If IsValid() Then                       ' Si los datos del formulario son Válidos
        CalMuestra                          ' Calcula la muestra antes de salir
        Me.Tag = EJECUTAR                   ' Se indica que el comando es EJECUTAR
                                            ' en la etiqueta del formulario
        Me.Hide                             ' Se oculta el fromulario
    Else
                                            ' Se obtienen los errores del formulario
                                            ' en la variable m_sMensaje
        strMensaje = MensajeValidacion
        Call MsgBox(strMensaje, vbExclamation Or vbSystemModal, Application.Name)
    End If
    
   On Error GoTo 0
   Exit Sub

cmdEjecutar_Click_Error:

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "frmCalTiempoMedio.cmdEjecutar_Click")
    '   Informa del error
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
End Sub

'-----------------------------|Funciones|----------------------------------------------*
'
'
'---------------------------------------------------------------------------------------
' Procedure : IsValid
' Author    : Charly
' Date      : 04/02/2018
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function IsValid() As Boolean
    Dim bNum As Numero
    Dim bInfo As New InfoSorteo
    
  On Error GoTo IsValid_Error
    '
    '   Inicializamos variables de control
    '
     mErrorValidacion = 0
     mCombinacion.Clear
    '
    '   Comprobamos el tipo de dato de cada control
    '   Segun el tipo de proceso
    '
    Select Case mTipoProceso
    Case 1:                         'Todos los Numeros
        '
        '   La fecha de pronostico no es valida
        '
        If Not IsDate(txtFePronostico.Text) Then
            mErrorValidacion = mErrorValidacion + ERR_FECHANOVALIDA
        Else
            mFechaSorteo = CDate(txtFePronostico.Text)
        End If
        '
        '
        '
    Case 2:                        ' Fecha de sorteo
        '
        '   La fecha de pronostico no es valida
        '
        If Not IsDate(txtFePronostico.Text) Then
            mErrorValidacion = mErrorValidacion + ERR_FECHANOVALIDA
        Else
            mFechaSorteo = CDate(txtFePronostico.Text)
        End If
        '
        '   La fecha no es de un sorteo
        '
        If Not bInfo.EsFechaSorteo(mFechaSorteo) Then
            mErrorValidacion = mErrorValidacion + ERR_FECHANOSORTEO
        End If
        '
        '   La fecha no es de un sorteo de la base de datos
        '
        If mFechaSorteo > mFinDataBase Then
            mErrorValidacion = mErrorValidacion + ERR_FECHANOSORTEO
        End If
        '
        '
        '
    Case 3:                         ' un conjuto de Numeros
        '
        '   Tenemos contenido en el primer Numero
        '
        If Len(txtNumero1.Text) > 0 Then
            '
            '   Si el Numero no es numerico
            '
            If Not IsNumeric(txtNumero1) Then
                mErrorValidacion = mErrorValidacion + ERR_NOESNumero
            '
            '   Si el Numero no esta vacio
            '
            Else
                '
                '   Creamos el objeto Numero
                '
                Set bNum = New Numero
                bNum.Valor = CInt(txtNumero1.Text)
                '
                '   Si es valido
                '
                If bNum.EsValido(JUEGO_DEFECTO) Then
                    '
                    '   Lo añadimos a la colección
                    '
                    mCombinacion.Add bNum
                Else
                    mErrorValidacion = mErrorValidacion + ERR_NumeroNOVALIDO
                End If
            End If
        End If
        '
        '   Segundo número
        '
        If Len(txtNumero2.Text) > 0 Then
            If Not IsNumeric(txtNumero2.Text) Then
                mErrorValidacion = mErrorValidacion + ERR_NOESNumero
            Else
                Set bNum = New Numero
                bNum.Valor = CInt(txtNumero2.Text)
                If bNum.EsValido(JUEGO_DEFECTO) Then
                    mCombinacion.Add bNum
                Else
                    mErrorValidacion = mErrorValidacion + ERR_NumeroNOVALIDO
                End If
            End If
        End If
        '
        '   Tercer número
        '
        If Len(txtNumero3.Text) > 0 Then
            If Not IsNumeric(txtNumero3.Text) Then
                mErrorValidacion = mErrorValidacion + ERR_NOESNumero
            Else
                Set bNum = New Numero
                bNum.Valor = CInt(txtNumero3.Text)
                If bNum.EsValido(JUEGO_DEFECTO) Then
                    mCombinacion.Add bNum
                Else
                    mErrorValidacion = mErrorValidacion + ERR_NumeroNOVALIDO
                End If
            End If
        End If
        '
        '   Si no hay Numeros
        '
        If mCombinacion.Count = 0 Then
            mErrorValidacion = mErrorValidacion + ERR_NumeroNOEXISTE
        End If
    End Select
    '
    '   Si el combo tiene fechas personalizadas comprueba la fecha
    '
    If mPeriodo.Tipo_Fecha = ctPersonalizadas Then
        If Not IsDate(txtFechaInicial.Text) Then
            mErrorValidacion = mErrorValidacion + ERR_FEFINNOVALIDA
        End If
        If Not IsDate(txtFechaFinal.Text) Then
            mErrorValidacion = mErrorValidacion + ERR_FEFINNOVALIDA
        End If
    End If
    '
    '   Configuramos los parametros de la muestra, si no hay errores
    '   previos
    '
    If mErrorValidacion = 0 Then
         With mParMuestra
            .Juego = JUEGO_DEFECTO
            .FechaAnalisis = mFechaSorteo
            .FechaFinal = mPeriodo.FechaFinal
            .FechaInicial = mPeriodo.FechaInicial
        End With
        '
        '   Comprobamos si es válido
        '
        If Not mParMuestra.Validar Then
            mErrorValidacion = mErrorValidacion + ERR_PARMUESTRAMAL
        End If
    End If
    '
    '   Evaluamos el código de error
    '
    If mErrorValidacion = 0 Then
        IsValid = True
    Else
        IsValid = False
    End If

   On Error GoTo 0
   Exit Function

IsValid_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmCalTiempoMedio.IsValid")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'---------------------------------------------------------------------------------------
' Procedure : MensajeValidacion
' Author    : Charly
' Date      : 04/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function MensajeValidacion() As String
    Dim sResult             As String
    
  On Error GoTo MensajeValidacion_Error
    '
    '   Si no hay error inicializamos el mensaje
    '
    If mErrorValidacion = 0 Then
        sResult = ""
    Else
        '
        '   Si hay algun error inicializamos la cabecera del error
        '
        sResult = MSG_ERRORESVALIDACION
    End If
    
    If (mErrorValidacion And ERR_NOESNumero) Then
        sResult = sResult & vbTab & MSG_NOESUNNumero & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_NumeroNOVALIDO) Then
        sResult = sResult & vbTab & MSG_NumeroNOVALIDO & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_NumeroNOEXISTE) Then
        sResult = sResult & vbTab & MSG_NumeroNOEXISTE & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_FECHANOVALIDA) Then
        sResult = sResult & vbTab & MSG_FECHANOVALIDA & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_FECHANOSORTEO) Then
        sResult = sResult & vbTab & MSG_FECHANOSORTEO & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_PARMUESTRAMAL) Then
        sResult = sResult & vbTab & MSG_PARMUESTRAMAL & mParMuestra.GetMensaje & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_FEININOVALIDA) Then
        sResult = sResult & vbTab & MSG_FEININOVALIDA & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_FEFINNOVALIDA) Then
        sResult = sResult & vbTab & MSG_FEFINNOVALIDA & vbCrLf
    End If
    '
    '   Devolvemos el mensaje
    '
    MensajeValidacion = sResult

   On Error GoTo 0
   Exit Function

MensajeValidacion_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmCalTiempoMedio.MensajeValidacion")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

'---------------------------------------------------------------------------------------
' Procedure : VisualizaControles
' Author    : Charly
' Date      : 06/02/2018
' Purpose   : Asigna valores a los controles y visualiza
'---------------------------------------------------------------------------------------
'
Private Sub VisualizaControles()
    Dim mInfo As New InfoSorteo         ' Información del sorteo
    
  On Error GoTo VisualizaControles_Error
    '
    '   Evaluar coherencia de datos
    '
    '
    If Not mInfo.EsFechaSorteo(mFechaSorteo) Then
        mFechaSorteo = mInfo.GetProximoSorteo(mFechaSorteo)
    End If
    '   Si la fecha de análisis no es de un sorteo
    '   vamos al siguiente
    '
    If mTipoProceso = 2 Then
        If mInfo.EsFechaSorteo(mFechaSorteo) Then
            mPeriodo.FechaFinal = mInfo.GetAnteriorSorteo(mFechaSorteo)
        End If
    End If
    '
    '   Si la fecha final del periodo no es de un sorteo
    '   salta a la fecha del sorteo anterior
    '
    If mInfo.EsFechaSorteo(mPeriodo.FechaFinal) Then
        mPeriodo.FechaFinal = mInfo.GetAnteriorSorteo(mPeriodo.FechaFinal)
    End If
    '
    '   Si la fecha inicial del periodo no es de un sorteo
    '   salta a la fecha del sorteo anterior
    '
    If mInfo.EsFechaSorteo(mPeriodo.FechaInicial) Then
        mPeriodo.FechaInicial = mInfo.GetAnteriorSorteo(mPeriodo.FechaInicial)
    End If
    
    If mPeriodo.FechaFinal > mFinDataBase Then
        mPeriodo.FechaFinal = mFinDataBase
    End If
    If mPeriodo.FechaInicial < mIniDataBase Then
        mPeriodo.FechaInicial = mIniDataBase
    End If
    
    
    '
    '   Configura el formulario para el tipo de proceso
    '
    Select Case mTipoProceso
        Case 1:                             ' Todos los número
            chkTodosNumeros.Value = True
            chkNumeros.Value = False
            txtNumero1.Enabled = False
            txtNumero2.Enabled = False
            txtNumero3.Enabled = False
            chkSorteo.Value = False
            lblFechaPronostico.Caption = "Fecha de Pronóstico"
            
        Case 2:                         ' Numeros de un sorteo
            chkTodosNumeros.Value = False
            chkNumeros.Value = False
            txtNumero1.Enabled = False
            txtNumero2.Enabled = False
            txtNumero3.Enabled = False
            chkSorteo.Value = True
            lblFechaPronostico.Caption = "Fecha de Sorteo"
        
        Case 3:                          ' uno o varios números
            chkTodosNumeros.Value = False
            chkNumeros.Value = True
            txtNumero1.Enabled = True
            txtNumero2.Enabled = True
            txtNumero3.Enabled = True
            chkSorteo.Value = False
            lblFechaPronostico.Caption = "Fecha de Pronóstico"
        
    End Select
    '
    '   Formateamos la fecha de sorteo
    '
    txtFePronostico.Text = Format(mFechaSorteo, "dd/mm/yyyy")
    '
    '   Formatea la fecha inicial y la coloca en el correspondiente caja de texto
    '
    txtFechaInicial.Text = Format(mPeriodo.FechaInicial, "dd/mm/yyyy")
    '
    '   Formatea la fecha final y la coloca en el correspondiente caja de texto
    '
    txtFechaFinal.Text = Format(mPeriodo.FechaFinal, "dd/mm/yyyy")
    '
    '   Si son fecha predefinidas bloqueamos los controles
    '
    If mPeriodo.Tipo_Fecha <> ctPersonalizadas Then
        '
        '   Desactiva los controles
        '
        txtFechaInicial.Enabled = False
        txtFechaFinal.Enabled = False
    Else
        '
        '   Si la fecha es personalizada activa los controles
        '
        txtFechaInicial.Enabled = True
        txtFechaFinal.Enabled = True
    End If
On Error GoTo 0
   Exit Sub

VisualizaControles_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "frmCalTiempoMedio.VisualizaControles")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CalMuestra
' Author    : Charly
' Date      : 14/02/2018
' Purpose   : Calcula la muestra estadistica
'---------------------------------------------------------------------------------------
'
Private Sub CalMuestra()
    Dim oRange As Range
    
  On Error GoTo CalMuestra_Error
    '
    '   Asignamos valores
    '
    With mParMuestra
        .Juego = JUEGO_DEFECTO
        .FechaAnalisis = mFechaSorteo
        .FechaFinal = mPeriodo.FechaFinal
        .FechaInicial = mPeriodo.FechaInicial
    End With
    '
    '   Si los parametros de la muestra están mal mensaje de error
    '
    If Not mParMuestra.Validar Then
        Err.Raise ERR_PARMUESTRAMAL, "CalMuestra", mParMuestra.GetMensaje
    End If
    '
    '       Calcula la Muestra
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set oRange = mDB.Resultados_Fechas(mParMuestra.FechaInicial, _
                                       mParMuestra.FechaFinal)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set mMuestra.ParametrosMuestra = mParMuestra
    '
    '   Calcula las bolas para este rango
    '
    mMuestra.Constructor oRange, JUEGO_DEFECTO
  
On Error GoTo 0
   Exit Sub

CalMuestra_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "frmCalTiempoMedio.CalMuestra")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------EOF| frmCalTiempoMedio.frm
