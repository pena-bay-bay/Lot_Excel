VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetodoOptimo 
   Caption         =   "Parámetros de Cálculo del Método Óptimo"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   OleObjectBlob   =   "frmMetodoOptimo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMetodoOptimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'---------------------------------------------------------------------------------------
' Module    : frmMetodoOptimo
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:02:01
' Purpose   :
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Base 0

'*-----------------| VARIABLES   |------------------------------------------------------
Dim m_periodo As New Periodo        ' Objeto que facilita el manejo de dos fechas
Dim m_ini_DataBase As Date          ' Fecha inicial de la base de datos
Dim m_fin_DataBase As Date          ' Fecha final de la base de datos
Dim DB As New BdDatos               ' Base de datos
Private mParametros As ParametrosSimulacion
Private mMetodo As ParametrosMetodoOld
'---------------------------------------------------------------------------------------
' Procedure : cboPeriodo_Change
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:02:56
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cboPeriodo_Change()
   On Error GoTo cboPeriodo_Change_Error
            
    m_periodo.Tipo_Fecha = cboPeriodo.ListIndex     ' Actualizamos el tipo de periodo
                                                    ' seleccionado
    If m_periodo.Tipo_Fecha <> 0 Then               ' Si el periodo no es personalizado
                                                    ' Si la fecha inicial no es menor
                                                    ' que la base de datos
        If (m_periodo.FechaInicial < m_ini_DataBase) Then
                                                    ' Selecciona la fecha menor de la
                                                    ' base de datos
            txtFechaInicial.Text = Format(m_ini_DataBase, "dd/mm/yyyy")
        Else
                                                    ' Formatea la fecha inicial y la
                                                    ' la coloca en la correspondiente caja
                                                    ' de texto
            txtFechaInicial.Text = Format(m_periodo.FechaInicial, "dd/mm/yyyy")
        End If
                                                    ' Si la fecha final es superior a
                                                    ' la fecha de la base de datos
        If (m_periodo.FechaFinal > m_fin_DataBase) Then
                                                    ' Selecciona la fecha mayor de la base
                                                    ' de datos
            txtFechaFinal.Text = Format(m_fin_DataBase, "dd/mm/yyyy")
        Else
                                                    ' Formatea la fecha final y la
                                                    ' coloca en el correspondiente
                                                    ' caja de texto
            txtFechaFinal.Text = Format(m_periodo.FechaFinal, "dd/mm/yyyy")
        End If
        txtFechaInicial.Enabled = False
        txtFechaFinal.Enabled = False
    Else
        txtFechaInicial.Enabled = True
        txtFechaFinal.Enabled = True
    End If
            
   On Error GoTo 0
       Exit Sub
            
cboPeriodo_Change_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoOptimo.cboPeriodo_Change", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "frmMetodoOptimo.cboPeriodo_Change", ErrDescription

End Sub
'---------------------------------------------------------------------------------------
' Procedure : cmdCancelar_Click
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:03:08
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR               ' Asigna en la etiqueta del formulario
                                        ' la clave de cerrar
    Me.Hide                             ' Oculta el formulario
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cmdEjecutar_Click
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:03:15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdEjecutar_Click()
    Dim bol_opcion_valida As Boolean
    Dim str_mensaje As String
    
   On Error GoTo cmdEjecutar_Click_Error

    bol_opcion_valida = False               ' Asigna a la variable de control que
                                            ' las opciones son erroneas
                                            ' Comprueba que el contenido del TextBox
                                            ' sea una fecha valida
    If (Not IsDate(Me.txtFechaInicial)) Then
                                            ' Si no es asi se carga el mensaje de error
                                            ' especifico de la fecha inicial
            str_mensaje = "No ha introducido una fecha inicial válida." + vbCrLf _
            + "Introduzca una fecha dentro del rango de los resultados."
            Me.txtFechaInicial.SetFocus     ' Focaliza el texbox del error
                                            ' Comprueba que el contenido del texBox de
                                            ' la fecha final es valida
    ElseIf (Not IsDate(Me.txtFechaFinal)) Then
                                            ' Si no lo es asigna el error del mensa
            str_mensaje = "No ha introducido una fecha final válida." + vbCrLf _
            + "Introduzca una fecha dentro del rango de los resultados."
            Me.txtFechaFinal.SetFocus       ' Focaliza el texbox del error
    ElseIf (Not IsNumeric(Me.txtDiasMuestra)) Then
                                            ' Si no lo es asigna el error del mensa
            str_mensaje = "Los días de la muestra debe ser numérico."
            Me.txtDiasMuestra.SetFocus       ' Focaliza el texbox del error
    
    ElseIf (Not IsNumeric(Me.txtDiasRetardo)) Then
                                                ' Si no lo es asigna el error del mensa
            str_mensaje = "Los días de retardo debe ser numérico."
            Me.txtDiasRetardo.SetFocus       ' Focaliza el texbox del error
    Else
    '
    '   Seguir con el resto de datos
        mParametros.ArrayMetodos = ContarMetodos
        mParametros.RangoAnalisis.Tipo_Fecha = cboPeriodo.ListIndex
        mParametros.RangoAnalisis.FechaInicial = CDate(txtFechaInicial.Text)
        mParametros.RangoAnalisis.FechaFinal = CDate(txtFechaFinal.Text)
        mMetodo.DiasMuestra = CInt(txtDiasMuestra.Text)
        mMetodo.DiasRetardo = CInt(txtDiasRetardo.Text)
        Me.Tag = EJECUTAR                   ' Si los datos son correctos se asigna
                                            ' la etiqueta EJECUTAR al formulario
        bol_opcion_valida = True            ' Se asigna true al control de opciones
                                            ' correctas
        SaveParametros
    End If
    
    If bol_opcion_valida Then               ' Si las opcines son correctas
        Me.Hide                             ' Se oculta el fromulario
    Else
                                            ' Si no lo son se emite un mensaje
                                            ' con el error
        MsgBox str_mensaje, vbExclamation + vbOKOnly, Me.Caption
    End If

   On Error GoTo 0
   Exit Sub

cmdEjecutar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEjecutar_Click of Formulario frmProbDosNumeros"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ContarMetodos
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:03:22
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ContarMetodos() As Variant
    Dim i As Integer, j As Integer, vMtd() As Variant
   On Error GoTo ContarMetodos_Error
            
     j = 0
     For i = 0 To lstMetodos.ListCount - 1
            If (lstMetodos.Selected(i) = True) Then
                j = j + 1
            End If
     Next i
     If (j = 0) Then
        ReDim vMtd(8)
        vMtd = Array(0, 1, 2, 3, 4, 5, 6, 7)
     Else
        ReDim vMtd(j - 1)
        j = 0
        For i = 0 To lstMetodos.ListCount - 1
            If (lstMetodos.Selected(i) = True) Then
                vMtd(j) = i
                j = j + 1
            End If
        Next i
     End If
     
     ContarMetodos = vMtd
            
   On Error GoTo 0
       Exit Function
            
ContarMetodos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoOptimo.ContarMetodos", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "frmMetodoOptimo.ContarMetodos", ErrDescription
End Function
'---------------------------------------------------------------------------------------
' Procedure : Parametros
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:03:34
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Parametros() As ParametrosSimulacion
    Set Parametros = mParametros
End Property
Public Property Set Parametros(ByVal vNewValue As ParametrosSimulacion)
    Set mParametros = vNewValue
End Property
'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:04:02
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    Dim mMetodoA As MetodoOld
   On Error GoTo UserForm_Initialize_Error
            
    Set mParametros = New ParametrosSimulacion
    Set mMetodo = mParametros.GetNewMetodo
    mMetodo.Id = 1
    mMetodo.DiasMuestra = 21
    mMetodo.DiasRetardo = 7
    mMetodo.Ordenacion = ctAleatorio
    mParametros.Add mMetodo
    m_ini_DataBase = DB.PrimerResultado         ' Obtiene la menor de las fechas de la
                                                ' base de datos
    m_fin_DataBase = DB.UltimoResultado         ' Obtiene la mayor de las fechas de la
                                                ' base de datos
                                                ' Formatea y asigna la fecha inicial
    txtFechaInicial.Text = Format(m_ini_DataBase, "dd/mm/yyyy")
                                                ' Formatea y asigna la fecha final
    txtFechaFinal.Text = Format(m_fin_DataBase, "dd/mm/yyyy")
                                                
    m_periodo.CargaTabla cboPeriodo             ' Carga el combo con los periodos
                                                ' predefinidos
                                            
    LoadParametros
    txtDiasMuestra.Text = CStr(mMetodo.DiasMuestra)
    txtDiasRetardo.Text = CStr(mMetodo.DiasRetardo)
    cboPeriodo.ListIndex = mParametros.RangoAnalisis.Tipo_Fecha
    
    Set mMetodoA = New MetodoOld
    mMetodoA.CargaTabla Me.lstMetodos
    lstMetodos.Selected(2) = True
    lstMetodos.Selected(3) = True
    lstMetodos.Selected(4) = True
    lstMetodos.Selected(5) = True
    lstMetodos.Selected(6) = True
    lstMetodos.Selected(7) = True
    lstMetodos.Selected(8) = True
    lstMetodos.Selected(9) = True
            
   On Error GoTo 0
       Exit Sub
            
UserForm_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoOptimo.UserForm_Initialize", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "frmMetodoOptimo.UserForm_Initialize", ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveParametros
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:05:11
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveParametros()
 
   On Error GoTo SaveParametros_Error
            
    If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
        THISLIBRO = ActiveWorkbook.Name
    End If
    
    With Workbooks(THISLIBRO).Worksheets("Variables")
        .Range("A9").Value = "Rango Comprobación"
        .Range("A10").Value = "Fecha Inicio"
        .Range("A11").Value = "Fecha Fin"
        .Range("A12").Value = "Dias Muestra"
        .Range("A13").Value = "Pronosticos"
        
        .Range("B9").Value = mParametros.RangoAnalisis.Tipo_Fecha
        .Range("C9").Value = mParametros.RangoAnalisis.Texto
        .Range("B10").Value = mParametros.RangoAnalisis.FechaInicial
        .Range("B11").Value = mParametros.RangoAnalisis.FechaFinal
        .Range("B12").Value = mMetodo.DiasMuestra
        .Range("B13").Value = mMetodo.DiasRetardo
        
    End With
            
   On Error GoTo 0
       Exit Sub
            
SaveParametros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoOptimo.SaveParametros", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "frmMetodoOptimo.SaveParametros", ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadParametros
' Author    : CHARLY
' Date      : jue, 18/sep/2014 00:05:16
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadParametros()
   On Error GoTo LoadParametros_Error
            
   If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
        THISLIBRO = ActiveWorkbook.Name
    End If

    With Workbooks(THISLIBRO).Worksheets("Variables")
        mParametros.RangoAnalisis.Tipo_Fecha = .Range("B9").Value
        If (mParametros.RangoAnalisis.Tipo_Fecha = ctPersonalizadas) Then
            mParametros.RangoAnalisis.FechaInicial = .Range("B10").Value
            mParametros.RangoAnalisis.FechaFinal = .Range("B11").Value
        End If
        mMetodo.DiasMuestra = .Range("B12").Value
        mMetodo.DiasRetardo = .Range("B13").Value
    End With
                
   On Error GoTo 0
       Exit Sub
            
LoadParametros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoOptimo.LoadParametros", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "frmMetodoOptimo.LoadParametros", ErrDescription
 
End Sub


