VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRealizarColoreado 
   Caption         =   "Colorear Resultados"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   OleObjectBlob   =   "frmRealizarColoreado.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmRealizarColoreado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Modulo    : frmRealizarColoreado
' Creado    : 16/03/2007  21:54
' Autor     : Carlos Almela Baeza
' Version   : 1.0.1 20/03/2007 9:49
' Objeto    : Formulario de captura de criterios para colorear los resultados
'---------------------------------------------------------------------------------------

Public Property Get Fecha_Sorteo() As Date
    If (IsDate(txtFecha)) Then
        Fecha_Sorteo = CDate(txtFecha)
    Else
        Fecha_Sorteo = Date
    End If
End Property

Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

Private Sub cmdEjecutar_Click()
    Dim bol_opcion_valida As Boolean
    Dim str_mensaje As String
    Dim objApuesta As Apuesta
    
    bol_opcion_valida = False
    
    Select Case MultiPage1.Value
    Case 0:
            If (Me.Tipo_Caracteristica = 0) Then
                str_mensaje = "No ha escogido ningún tipo de caracteristicas " + vbCrLf _
                + "para colorear los Numeros."
            Else
                Me.Tag = COLOREAR_CARACTERISTICAS
                bol_opcion_valida = True
            End If
    Case 1:
            Set objApuesta = TextCombinacion
            If Not (objApuesta Is Nothing) Then
                If (objApuesta.Pronosticos = 0) Then
                    str_mensaje = "No ha introducido ningún " + vbCrLf _
                    + "número válido para colorear."
                Else
                    Me.Tag = COLOREAR_NumeroS
                    bol_opcion_valida = True
                End If
            Else
                Exit Sub
            End If
    Case 2:
            If (Not IsDate(Me.txtFecha)) Then
                str_mensaje = "No ha introducido una fecha válida." + vbCrLf _
                + "Introduzca una fecha dentro del rango de los resultados."
            Else
                Me.Tag = COLOREAR_UNAFECHA
                bol_opcion_valida = True
            End If
    End Select
    If bol_opcion_valida Then
        Me.Hide
    Else
        MsgBox str_mensaje, vbExclamation + vbOKOnly, Me.Caption
    End If
End Sub


Public Property Get Tipo_Caracteristica() As Variant
    Select Case True
        Case optSeleccion1: Tipo_Caracteristica = 1     ' Pares e impares
        Case optSeleccion2: Tipo_Caracteristica = 2     ' Altos y Bajos
        Case optSeleccion3: Tipo_Caracteristica = 3     ' Decenas
        Case optSeleccion4: Tipo_Caracteristica = 4     ' Terminaciones
        Case optSeleccion5: Tipo_Caracteristica = 5     ' Consecutivos
        Case Else: Tipo_Caracteristica = 0              ' No se ha escogino ninguno
    End Select
End Property


Public Property Get TextCombinacion() As Apuesta
    Dim obj_apuesta     As Apuesta
    Dim objNumero       As Numero
    Dim a_Numeros(10)   As Integer
    Dim i               As Integer
    Dim j               As Integer
    
On Error GoTo TextCombinacion_Error
    Set obj_apuesta = New Apuesta
    i = 0
    If (IsNumeric(txt_N1.Text)) Then
        a_Numeros(i) = CInt(txt_N1.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N2.Text)) Then
        a_Numeros(i) = CInt(txt_N2.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N3.Text)) Then
        a_Numeros(i) = CInt(txt_N3.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N4.Text)) Then
        a_Numeros(i) = CInt(txt_N4.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N5.Text)) Then
        a_Numeros(i) = CInt(txt_N5.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N6.Text)) Then
        a_Numeros(i) = CInt(txt_N6.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N7.Text)) Then
        a_Numeros(i) = CInt(txt_N7.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N8.Text)) Then
        a_Numeros(i) = CInt(txt_N8.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N9.Text)) Then
        a_Numeros(i) = CInt(txt_N9.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N10.Text)) Then
        a_Numeros(i) = CInt(txt_N10.Text)
        i = i + 1
    End If
    If (IsNumeric(txt_N11.Text)) Then
        a_Numeros(i) = CInt(txt_N11.Text)
        i = i + 1
    End If
    j = i - 1
    For i = 0 To j
        If (a_Numeros(i) > 0 And a_Numeros(i) < 50) Then
            Set objNumero = New Numero
            objNumero.Valor = a_Numeros(i)
            obj_apuesta.Combinacion.Add objNumero
        End If
    Next i
    Set TextCombinacion = obj_apuesta
    
   On Error GoTo 0
   Exit Sub
   
TextCombinacion_Error:
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Informa del error
   Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    
End Property

