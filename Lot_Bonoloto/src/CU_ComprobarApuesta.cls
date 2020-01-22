VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CU_ComprobarApuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : CU_ComprobarApuesta
' *
' *     Tipo       : Módulo de clase
' *     Autor      : CAB3780Y
' *     Creacion   : SSS, 12/ago/2007 19:11
' *     Version    : 1.0
' *     Asunto     : Caso de uso que realiza la tarea de comprobar si una apuesta
' *                  tiene aciertos y determina el premio conseguido
' *
' *============================================================================*
Option Explicit
Public MyApuesta As Apuesta
Public Sorteo As Sorteo
Public Premio As Premio

Private m_apta As Apuesta
Private m_rslt As Resultado
Private m_Complementario As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 12/ago/2007 19:12
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    'Set m_apta = New Apuesta
    'Set m_rslt = New Resultado
    Set MyApuesta = New Apuesta
    Set Sorteo = New Sorteo
    Set Premio = New Premio
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' DateTime  : 12/ago/2007 19:12
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set m_apta = Nothing
    Set m_rslt = Nothing
    Set MyApuesta = Nothing
    Set Sorteo = Nothing
    Set Premio = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Apuesta
' DateTime  : 12/ago/2007 19:12
' Author    : Carlos Almela Baeza
' Purpose   : precated
'---------------------------------------------------------------------------------------
'
Public Property Get Apuesta() As Apuesta
    Set Apuesta = m_apta
End Property
Public Property Set Apuesta(objApuesta As Apuesta)
    Set m_apta = objApuesta
End Property

'---------------------------------------------------------------------------------------
' Procedure : Resultado
' DateTime  : 12/ago/2007 19:12
' Author    : Carlos Almela Baeza
' Purpose   : precated
'---------------------------------------------------------------------------------------
'
Public Property Get Resultado() As Resultado
    Set Resultado = m_rslt
End Property
Public Property Set Resultado(objResultado As Resultado)
    Set m_rslt = objResultado
End Property

'---------------------------------------------------------------------------------------
' Procedure : Get_Aciertos
' DateTime  : 12/ago/2007 19:12
' Author    : Carlos Almela Baeza
' Purpose   : precated
'---------------------------------------------------------------------------------------
'
Public Function Get_Aciertos(Optional Mas_Complementario As Boolean = True) As Integer
    Dim m_aciertos As Integer
    Dim m_array As Variant
    Dim K As Integer
    Dim n As Variant
    
    If m_apta Is Nothing Then Err.Raise "falta la apuesta"
    If m_rslt Is Nothing Then Err.Raise "falta el resultado"
    m_array = m_apta.Pronosticos
    m_aciertos = 0
    m_Complementario = False
    For Each n In m_array
        K = n
        If (m_rslt.Existe(K, Mas_Complementario)) Then
            m_aciertos = m_aciertos + 1
        End If
        If K = m_rslt.Complementario Then
            m_Complementario = True
        End If
    Next n
    Get_Aciertos = m_aciertos
End Function


' *============================================================================*
' *     Procedimiento       : Get_Premio ( Function )
' *     Version             : 1.0 lun, 09/06/2008 08:57
' *     Autor               : CAB3780Y
' *     Retorno             : String
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            : precated
' *
' *============================================================================*
Public Function Get_Premio() As String
    Dim mPremio As String
    Dim nAciertos As Integer
    nAciertos = Get_Aciertos(False)
    If (m_rslt.Dia = "D") Then
        Select Case nAciertos
            Case 2: mPremio = "8ª"
            Case 3: mPremio = "6ª"
            Case 4: mPremio = "4ª"
            Case 5: mPremio = "2ª"
        End Select
    Else
        If m_apta.Pronosticos = 5 Then
            Select Case nAciertos
                Case 2: mPremio = "5ª"
                Case 3: mPremio = "4ª"
                Case 4:
                    If (m_Complementario) Then
                        mPremio = "3ª"
                    Else
                        mPremio = "2ª"
                    End If
                Case 5: mPremio = "1ª"
            End Select
        Else
            Select Case nAciertos
                Case 3: mPremio = "5ª"
                Case 4: mPremio = "4ª"
                Case 5:
                    If (m_Complementario) Then
                        mPremio = "3ª"
                    Else
                        mPremio = "2ª"
                    End If
                Case 6: mPremio = "1ª"
            End Select
        End If
    End If
    Get_Premio = mPremio
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetImporteEsperado
' Author    : CHARLY
' Date      : 08/02/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetImporteEsperado() As Currency
    Dim oPremio            As New Premio
    Dim nAciertos          As Integer
   On Error GoTo GetImporteEsperado_Error

    nAciertos = Get_Aciertos(False)
    oPremio.FechaSorteo = Me.Resultado.Fecha
    oPremio.ModalidadJuego = Me.Resultado.TipoJuego
    oPremio.Pronosticos = Me.Apuesta.Pronosticos
    Select Case nAciertos
        Case 0: oPremio.NumerosAcertados = SinAciertos
        Case 1: oPremio.NumerosAcertados = SinAciertos
        Case 2: oPremio.NumerosAcertados = SinAciertos
            If oPremio.Pronosticos = 5 Then
                oPremio.NumerosAcertados = DosAciertos
            End If
        Case 3: oPremio.NumerosAcertados = TresAciertos
        Case 4: oPremio.NumerosAcertados = CuatroAciertos
            If oPremio.Pronosticos = 5 And m_Complementario Then
                oPremio.NumerosAcertados = CuatroAciertosMasC
            End If
        Case 5: oPremio.NumerosAcertados = CincoAciertos
            If m_Complementario Then
                oPremio.NumerosAcertados = CincoAciertosMasC
            End If
        Case 6: oPremio.NumerosAcertados = SeisAciertos
            If m_Complementario Then
                oPremio.NumerosAcertados = SeisAciertosMasC
            End If
    End Select
    
    GetImporteEsperado = oPremio.GetPremioEsperado

   On Error GoTo 0
   Exit Function

GetImporteEsperado_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "CU_ComprobarApuesta.GetImporteEsperado")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function
' *============================================================================*
' *     Procedimiento       : Get_ImportePremio ( Function )
' *     Version             : 1.0 lun, 09/06/2008 09:01
' *     Autor               : CAB3780Y
' *     Retorno             : Currency
' *     Objetivo            : Asignar un importe mínimo a la apuesta
' *                           Bonoloto
' *============================================================================*
Public Function Get_ImportePremio() As Currency
        Dim mImporte As Currency
        Dim nAciertos As Integer
    On Error GoTo Get_ImportePremio_Error:
        nAciertos = Get_Aciertos(False)
        If (m_rslt.Dia = "D") Then
            Select Case nAciertos
                Case 2: mImporte = (3 * 4)
                Case 3: mImporte = (15 * 4) + (3 * 4)
                Case 4: mImporte = (200 * 4) + (15 * 4) + (3 * 4)
                Case 5: mImporte = 300000
            End Select
        Else
            If (m_rslt.Dia = "J") Or (m_rslt.Dia = "S") Then
                Select Case nAciertos
                    Case 3: mImporte = (8 * 4)
                    Case 4: mImporte = (85 * 3) + (8 * 4)
                    Case 5:
                        If (m_Complementario) Then
                            mImporte = 3000
                        Else
                            mImporte = 75000
                        End If
                    Case 6: mImporte = 1000000
                End Select
            Else
                Select Case nAciertos
                    Case 3: mImporte = (4 * 4)
                    Case 4: mImporte = (30 * 3) + (4 * 3)
                    Case 5:
                        If (m_Complementario) Then
                            mImporte = 1000
                        Else
                            mImporte = 100000
                        End If
                    Case 6: mImporte = 450000
                End Select
            End If
        End If
        Get_ImportePremio = mImporte
        
Get_ImportePremio_exit:
        Exit Function
        
Get_ImportePremio_Error:
        MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure cmd_Estadisticas of Módulo <mod>"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPremio
' Author    : Charly
' Date      : 16/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetPremio() As Premio
    Dim objResult   As Premio
    Dim n           As Numero
    Dim col         As Collection
    
  On Error GoTo GetPremio_Error
    
    Set objResult = New Premio
    objResult.FechaSorteo = Sorteo.Fecha
    objResult.ModalidadJuego = Sorteo.Juego
    objResult.Pronosticos = MyApuesta.Pronosticos
    Set col = MyApuesta.Combinacion.Numeros
    
    For Each n In col
    
        If Sorteo.Existe(n.Valor) Then
            objResult.BolasAcertadas = objResult.BolasAcertadas + 1
            If Sorteo.Complementario = n.Valor Then
                objResult.ComplementarioAcertado = True
            End If
        End If
        
    Next n
        
    Set GetPremio = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetPremio_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "CU_ComprobarApuesta.GetPremio")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

Private Function GetCategoria(datAciertos As Integer, datComple As Boolean) As CategoriaPremio

    Dim eResult As CategoriaPremio


    GetCategoria = eResult

End Function
